class AttendanceSystem {
    constructor() {
        this.rawData = null;
        this.processedData = null;
        this.reports = null;
        this.init();
    }

    init() {
        this.setupEventListeners();
    }

    setupEventListeners() {
        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('excelFile');
        const processBtn = document.getElementById('processBtn');
        const resetBtn = document.getElementById('resetBtn');
        const downloadRawBtn = document.getElementById('downloadRawBtn');
        const downloadReportBtn = document.getElementById('downloadReportBtn');

        uploadArea.addEventListener('click', () => fileInput.click());
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.style.borderColor = '#3498db';
            uploadArea.style.background = '#f0f8ff';
        });
        uploadArea.addEventListener('dragleave', () => {
            uploadArea.style.borderColor = '#dee2e6';
            uploadArea.style.background = 'white';
        });
        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.style.borderColor = '#dee2e6';
            uploadArea.style.background = 'white';
            
            const file = e.dataTransfer.files[0];
            if (file && this.isValidExcelFile(file)) {
                this.handleFile(file);
            } else {
                this.showToast('Harap upload file Excel (.xlsx, .xls)', 'error');
            }
        });

        fileInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (file && this.isValidExcelFile(file)) {
                this.handleFile(file);
            }
        });

        processBtn.addEventListener('click', () => this.processData());
        resetBtn.addEventListener('click', () => this.resetSystem());
        downloadRawBtn.addEventListener('click', () => this.downloadRawData());
        downloadReportBtn.addEventListener('click', () => this.downloadReport());
    }

    isValidExcelFile(file) {
        const validTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-excel'
        ];
        const validExtensions = ['.xlsx', '.xls'];
        
        return validTypes.includes(file.type) || 
               validExtensions.some(ext => file.name.toLowerCase().endsWith(ext));
    }

    handleFile(file) {
        this.showFileInfo(file);
        this.readExcelFile(file);
    }

    showFileInfo(file) {
        const fileInfo = document.getElementById('fileInfo');
        const fileName = document.getElementById('fileName');
        const fileSize = document.getElementById('fileSize');
        const fileType = document.getElementById('fileType');
        const processBtn = document.getElementById('processBtn');

        fileName.textContent = file.name;
        fileSize.textContent = this.formatFileSize(file.size);
        fileType.textContent = file.type || 'Excel file';
        
        fileInfo.classList.add('show');
        processBtn.disabled = false;
    }

    formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    readExcelFile(file) {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                this.rawData = workbook;
                this.showPreview(workbook);
            } catch (error) {
                this.showToast('Error membaca file Excel: ' + error.message, 'error');
            }
        };
        reader.readAsArrayBuffer(file);
    }

    showPreview(workbook) {
        const previewContent = document.getElementById('previewContent');
        const loadingPreview = document.getElementById('loadingPreview');
        
        loadingPreview.classList.add('show');
        
        setTimeout(() => {
            let html = '<div class="preview-tabs">';
            
            workbook.SheetNames.forEach((sheetName, index) => {
                const sheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
                
                html += `
                    <div class="sheet-tab">
                        <h4>Sheet: ${sheetName}</h4>
                        <div class="table-responsive">
                            <table class="preview-table">
                                <thead>
                                    <tr>
                `;
                
                if (jsonData.length > 0) {
                    const headers = jsonData[0];
                    headers.forEach(header => {
                        html += `<th>${header || ''}</th>`;
                    });
                }
                
                html += `
                    </tr>
                </thead>
                <tbody>
                `;
                
                for (let i = 1; i < Math.min(10, jsonData.length); i++) {
                    html += '<tr>';
                    jsonData[i].forEach(cell => {
                        html += `<td>${cell || ''}</td>`;
                    });
                    html += '</tr>';
                }
                
                if (jsonData.length > 10) {
                    html += `<tr><td colspan="${jsonData[0].length}" style="text-align: center; color: #666;">
                        ... dan ${jsonData.length - 10} baris lainnya
                    </td></tr>`;
                }
                
                html += `
                                </tbody>
                            </table>
                        </div>
                    </div>
                `;
            });
            
            html += '</div>';
            previewContent.innerHTML = html;
            loadingPreview.classList.remove('show');
            
            document.getElementById('rawSheets').textContent = workbook.SheetNames.length;
            
        }, 1000);
    }

    processData() {
        if (!this.rawData) {
            this.showToast('Harap upload file Excel terlebih dahulu', 'error');
            return;
        }

        const loadingPreview = document.getElementById('loadingPreview');
        const processBtn = document.getElementById('processBtn');
        
        loadingPreview.classList.add('show');
        processBtn.disabled = true;
        processBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Memproses...';

        setTimeout(() => {
            try {
                this.processedData = this.calculateAttendance(this.rawData);
                this.reports = this.generateReports(this.processedData);
                
                this.showSummary(this.processedData);
                this.showReportSection();
                
                this.showToast('Data berhasil diproses!', 'success');
                
            } catch (error) {
                this.showToast('Error memproses data: ' + error.message, 'error');
            } finally {
                loadingPreview.classList.remove('show');
                processBtn.disabled = false;
                processBtn.innerHTML = '<i class="fas fa-cogs"></i> Proses Data';
            }
        }, 2000);
    }

    calculateAttendance(workbook) {
        const result = {
            employees: [],
            attendance: [],
            overtime: [],
            summary: {
                totalEmployees: 0,
                totalWorkingDays: 0,
                totalWorkingHours: 0,
                totalOvertimeHours: 0,
                totalOvertimePay: 0
            }
        };

        // Process NOVEMBER sheet
        const novSheet = workbook.Sheets['NOVEMBER'];
        if (novSheet) {
            const jsonData = XLSX.utils.sheet_to_json(novSheet, { header: 1 });
            
            const employeeMap = new Map();
            const attendanceMap = new Map();
            
            // Get employee list from first column
            for (let i = 1; i < Math.min(15, jsonData.length); i++) {
                if (jsonData[i][1]) {
                    const employee = {
                        id: i,
                        name: jsonData[i][1].trim(),
                        position: this.getEmployeePosition(jsonData[i][1]),
                        department: this.getEmployeeDepartment(jsonData[i][1])
                    };
                    employeeMap.set(employee.name, employee);
                }
            }
            
            // Process attendance records (starting from row where column E has "Nama")
            let startRow = 0;
            for (let i = 0; i < jsonData.length; i++) {
                if (jsonData[i][4] === 'Nama') {
                    startRow = i + 1;
                    break;
                }
            }
            
            for (let i = startRow; i < jsonData.length; i++) {
                const name = jsonData[i][4];
                const timeStr = jsonData[i][5];
                
                if (name && timeStr) {
                    const { date, time, type } = this.parseDateTime(timeStr);
                    
                    const key = `${name}_${date}`;
                    if (!attendanceMap.has(key)) {
                        attendanceMap.set(key, {
                            employeeName: name,
                            date: date,
                            checkIn: null,
                            checkOut: null
                        });
                    }
                    
                    const record = attendanceMap.get(key);
                    if (type === 'in') {
                        record.checkIn = time;
                    } else {
                        record.checkOut = time;
                    }
                    
                    attendanceMap.set(key, record);
                }
            }
            
            result.employees = Array.from(employeeMap.values());
            result.attendance = Array.from(attendanceMap.values());
        }

        // Calculate summary
        result.summary.totalEmployees = result.employees.length;
        result.summary.totalWorkingDays = new Set(result.attendance.map(a => a.date)).size;
        
        let totalHours = 0;
        let totalOvertime = 0;
        
        result.attendance.forEach(record => {
            if (record.checkIn && record.checkOut) {
                const hours = this.calculateWorkingHours(record.checkIn, record.checkOut);
                totalHours += hours.total;
                
                if (hours.overtime > 0) {
                    totalOvertime += hours.overtime;
                    result.overtime.push({
                        employeeName: record.employeeName,
                        date: record.date,
                        overtimeHours: hours.overtime,
                        overtimePay: hours.overtime * 12500 // Rate lembur
                    });
                }
            }
        });
        
        result.summary.totalWorkingHours = Math.round(totalHours);
        result.summary.totalOvertimeHours = Math.round(totalOvertime);
        result.summary.totalOvertimePay = result.overtime.reduce((sum, ot) => sum + ot.overtimePay, 0);

        return result;
    }

    getEmployeePosition(name) {
        const positions = {
            'Windy': 'Staff',
            'Bu Ati': 'Kepala Bagian',
            'Pa Ardhi': 'Dosen',
            'Pak Ardhi': 'Dosen',
            'Pa Saji': 'Staff',
            'Pak Saji': 'Staff',
            'Pa Irvan': 'Dosen',
            'Pak Irvan': 'Dosen',
            'Bu Wahyu': 'Admin',
            'Bu Elzi': 'Staff',
            'Pa Nanang': 'Staff',
            'Nanang': 'Staff',
            'Devi': 'Admin',
            'Intan': 'Staff',
            'Alifah': 'Staff',
            'Pebi': 'Staff',
            'Dian': 'Staff',
            'Diana': 'Staff',
            'Rafly': 'Staff',
            'Erni': 'Staff'
        };
        return positions[name] || 'Staff';
    }

    getEmployeeDepartment(name) {
        const departments = {
            'Windy': 'Administrasi',
            'Bu Ati': 'Administrasi',
            'Pa Ardhi': 'Teknik Informatika',
            'Pak Ardhi': 'Teknik Informatika',
            'Pa Saji': 'Teknik Elektro',
            'Pak Saji': 'Teknik Elektro',
            'Pa Irvan': 'Teknik Mesin',
            'Pak Irvan': 'Teknik Mesin',
            'Bu Wahyu': 'Administrasi',
            'Bu Elzi': 'Administrasi',
            'Pa Nanang': 'Teknik Elektro',
            'Nanang': 'Teknik Elektro',
            'Devi': 'Administrasi',
            'Intan': 'Teknik Informatika',
            'Alifah': 'Staff',
            'Pebi': 'Staff',
            'Dian': 'Staff',
            'Diana': 'Staff',
            'Rafly': 'Staff',
            'Erni': 'Staff'
        };
        return departments[name] || 'Administrasi';
    }

    parseDateTime(dateTimeStr) {
        // Handle multiple date formats
        let date, time;
        
        if (dateTimeStr.includes('/')) {
            // Format: DD/MM/YYYY HH:MM
            const parts = dateTimeStr.split(' ');
            const dateParts = parts[0].split('/');
            date = `${dateParts[2]}-${dateParts[1]}-${dateParts[0]}`;
            time = parts[1];
        } else if (dateTimeStr.includes('-')) {
            // Format: YYYY-MM-DD HH:MM:SS
            const parts = dateTimeStr.split(' ');
            date = parts[0];
            time = parts[1].substring(0, 5);
        }
        
        const hour = parseInt(time.split(':')[0]);
        const type = hour < 12 ? 'in' : 'out';
        
        return { date, time, type };
    }

    calculateWorkingHours(checkIn, checkOut) {
        const [inHour, inMin] = checkIn.split(':').map(Number);
        const [outHour, outMin] = checkOut.split(':').map(Number);
        
        let startMinutes = inHour * 60 + inMin;
        let endMinutes = outHour * 60 + outMin;
        
        // Handle overnight shifts
        if (endMinutes < startMinutes) {
            endMinutes += 24 * 60;
        }
        
        const totalMinutes = endMinutes - startMinutes;
        const totalHours = totalMinutes / 60;
        
        // Normal work hours: 8 hours
        const normalHours = 8;
        const overtime = Math.max(0, totalHours - normalHours);
        
        return {
            total: totalHours,
            normal: Math.min(totalHours, normalHours),
            overtime: overtime
        };
    }

    showSummary(data) {
        document.getElementById('summaryCards').style.display = 'grid';
        
        document.getElementById('totalEmployees').textContent = data.summary.totalEmployees;
        document.getElementById('totalDays').textContent = data.summary.totalWorkingDays;
        document.getElementById('totalHours').textContent = data.summary.totalWorkingHours;
        document.getElementById('totalOvertime').textContent = data.summary.totalOvertimeHours;
    }

    showReportSection() {
        document.getElementById('reportSection').style.display = 'block';
    }

    generateReports(data) {
        const reports = {
            raw: this.createRawDataReport(data),
            calculated: this.createCalculatedReport(data)
        };
        
        return reports;
    }

    createRawDataReport(data) {
        const wb = XLSX.utils.book_new();
        
        // Sheet 1: Data Karyawan
        const employeeData = [
            ['DATA KARYAWAN'],
            ['ID', 'Nama', 'Jabatan', 'Departemen', 'Status']
        ];
        
        data.employees.forEach(emp => {
            employeeData.push([
                emp.id,
                emp.name,
                emp.position,
                emp.department,
                'Aktif'
            ]);
        });
        
        const ws1 = XLSX.utils.aoa_to_sheet(employeeData);
        XLSX.utils.book_append_sheet(wb, ws1, "Data_Karyawan");
        
        // Sheet 2: Presensi
        const attendanceData = [
            ['DATA PRESENSI'],
            ['Nama', 'Tanggal', 'Check In', 'Check Out', 'Durasi', 'Status']
        ];
        
        data.attendance.forEach(record => {
            const duration = record.checkIn && record.checkOut 
                ? this.calculateWorkingHours(record.checkIn, record.checkOut).total.toFixed(2) + ' jam'
                : '-';
            
            const status = this.getAttendanceStatus(record.checkIn, record.checkOut);
            
            attendanceData.push([
                record.employeeName,
                record.date,
                record.checkIn || '-',
                record.checkOut || '-',
                duration,
                status
            ]);
        });
        
        const ws2 = XLSX.utils.aoa_to_sheet(attendanceData);
        XLSX.utils.book_append_sheet(wb, ws2, "Data_Presensi");
        
        // Sheet 3: Lembur
        const overtimeData = [
            ['DATA LEMBUR'],
            ['Nama', 'Tanggal', 'Jam Lembur', 'Rate', 'Total Uang Lembur']
        ];
        
        data.overtime.forEach(ot => {
            overtimeData.push([
                ot.employeeName,
                ot.date,
                ot.overtimeHours.toFixed(2),
                'Rp 12.500/jam',
                'Rp ' + ot.overtimePay.toLocaleString('id-ID')
            ]);
        });
        
        const ws3 = XLSX.utils.aoa_to_sheet(overtimeData);
        XLSX.utils.book_append_sheet(wb, ws3, "Data_Lembur");
        
        return wb;
    }

    createCalculatedReport(data) {
        const wb = XLSX.utils.book_new();
        
        // Sheet 1: Laporan Presensi
        const reportData = [
            ['LAPORAN PRESENSI DAN LEMBUR'],
            ['Periode: November 2025'],
            [''],
            ['No', 'Nama', 'Jabatan', 'Total Hari', 'Total Jam', 'Jam Lembur', 'Uang Lembur', 'Total']
        ];
        
        data.employees.forEach((emp, index) => {
            const empAttendance = data.attendance.filter(a => a.employeeName === emp.name);
            const empOvertime = data.overtime.filter(ot => ot.employeeName === emp.name);
            
            const totalDays = new Set(empAttendance.map(a => a.date)).size;
            const totalHours = empAttendance.reduce((sum, a) => {
                if (a.checkIn && a.checkOut) {
                    return sum + this.calculateWorkingHours(a.checkIn, a.checkOut).total;
                }
                return sum;
            }, 0);
            
            const totalOvertime = empOvertime.reduce((sum, ot) => sum + ot.overtimeHours, 0);
            const totalPay = empOvertime.reduce((sum, ot) => sum + ot.overtimePay, 0);
            
            reportData.push([
                index + 1,
                emp.name,
                emp.position,
                totalDays,
                totalHours.toFixed(2),
                totalOvertime.toFixed(2),
                'Rp ' + totalPay.toLocaleString('id-ID'),
                'Rp ' + totalPay.toLocaleString('id-ID')
            ]);
        });
        
        // Summary row
        reportData.push(['']);
        reportData.push(['TOTAL', '', '', 
            data.summary.totalWorkingDays,
            data.summary.totalWorkingHours,
            data.summary.totalOvertimeHours,
            '',
            'Rp ' + data.summary.totalOvertimePay.toLocaleString('id-ID')
        ]);
        
        const ws1 = XLSX.utils.aoa_to_sheet(reportData);
        XLSX.utils.book_append_sheet(wb, ws1, "Laporan_Presensi");
        
        // Sheet 2: Detail Perhitungan
        const calculationData = [
            ['DETAIL PERHITUNGAN LEMBUR'],
            ['Rate Lembur: Rp 12.500 per jam'],
            [''],
            ['Nama', 'Tanggal', 'Jam Masuk', 'Jam Pulang', 'Total Jam', 'Jam Normal', 'Jam Lembur', 'Uang Lembur']
        ];
        
        data.attendance.forEach(record => {
            if (record.checkIn && record.checkOut) {
                const hours = this.calculateWorkingHours(record.checkIn, record.checkOut);
                const overtimePay = hours.overtime * 12500;
                
                calculationData.push([
                    record.employeeName,
                    record.date,
                    record.checkIn,
                    record.checkOut,
                    hours.total.toFixed(2),
                    hours.normal.toFixed(2),
                    hours.overtime.toFixed(2),
                    'Rp ' + overtimePay.toLocaleString('id-ID')
                ]);
            }
        });
        
        const ws2 = XLSX.utils.aoa_to_sheet(calculationData);
        XLSX.utils.book_append_sheet(wb, ws2, "Detail_Perhitungan");
        
        return wb;
    }

    getAttendanceStatus(checkIn, checkOut) {
        if (!checkIn) return 'Tidak Hadir';
        if (!checkOut) return 'Hadir (Belum Check Out)';
        
        const [hour] = checkIn.split(':').map(Number);
        if (hour > 9) return 'Terlambat';
        
        return 'Hadir';
    }

    downloadRawData() {
        if (!this.reports?.raw) {
            this.showToast('Tidak ada data untuk diunduh', 'error');
            return;
        }
        
        XLSX.writeFile(this.reports.raw, `DATA_PRESENSI_MENTAH_${new Date().toISOString().slice(0,10)}.xlsx`);
        this.showToast('Data presensi berhasil diunduh!', 'success');
    }

    downloadReport() {
        if (!this.reports?.calculated) {
            this.showToast('Tidak ada laporan untuk diunduh', 'error');
            return;
        }
        
        XLSX.writeFile(this.reports.calculated, `LAPORAN_PRESENSI_${new Date().toISOString().slice(0,10)}.xlsx`);
        this.showToast('Laporan berhasil diunduh!', 'success');
    }

    resetSystem() {
        this.rawData = null;
        this.processedData = null;
        this.reports = null;
        
        // Reset UI
        document.getElementById('fileInfo').classList.remove('show');
        document.getElementById('previewContent').innerHTML = '<p class="text-muted">Data akan muncul setelah file diupload dan diproses</p>';
        document.getElementById('processBtn').disabled = true;
        document.getElementById('excelFile').value = '';
        document.getElementById('summaryCards').style.display = 'none';
        document.getElementById('reportSection').style.display = 'none';
        
        this.showToast('Sistem telah direset', 'info');
    }

    showToast(message, type = 'info') {
        const toastContainer = document.getElementById('toastContainer');
        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        toast.innerHTML = `
            <i class="fas fa-${type === 'success' ? 'check-circle' : type === 'error' ? 'exclamation-circle' : 'info-circle'}"></i>
            <span>${message}</span>
        `;
        
        toastContainer.appendChild(toast);
        
        setTimeout(() => {
            toast.remove();
        }, 5000);
    }
}

// Initialize the system when page loads
window.addEventListener('DOMContentLoaded', () => {
    const attendanceSystem = new AttendanceSystem();
    window.attendanceSystem = attendanceSystem; // Make it available globally if needed
});
