// Main JavaScript file for the attendance processing system

document.addEventListener('DOMContentLoaded', function() {
    // DOM Elements
    const uploadZone = document.getElementById('upload-zone');
    const browseBtn = document.getElementById('browse-btn');
    const fileInput = document.getElementById('excel-file');
    const fileInfo = document.getElementById('file-info');
    const fileName = document.getElementById('file-name');
    const fileSize = document.getElementById('file-size');
    const progressBar = document.getElementById('progress-bar');
    const progressText = document.getElementById('progress-text');
    const cancelBtn = document.getElementById('cancel-btn');
    const processBtn = document.getElementById('process-btn');
    const resultsSection = document.getElementById('results-section');
    const originalTable = document.getElementById('original-table').getElementsByTagName('tbody')[0];
    const processedTable = document.getElementById('processed-table').getElementsByTagName('tbody')[0];
    
    // Statistics elements
    const totalEmployees = document.getElementById('total-employees');
    const totalDays = document.getElementById('total-days');
    const totalSalary = document.getElementById('total-salary');
    const totalOvertime = document.getElementById('total-overtime');
    
    // Download buttons
    const downloadOriginalBtn = document.getElementById('download-original');
    const downloadProcessedBtn = document.getElementById('download-processed');
    const downloadBothBtn = document.getElementById('download-both');
    
    // Tab switching
    const tabBtns = document.querySelectorAll('.tab-btn');
    
    // Data storage
    let originalData = [];
    let processedData = [];
    let currentFile = null;
    let uploadProgressInterval = null;
    
    // Event Listeners
    uploadZone.addEventListener('click', () => fileInput.click());
    browseBtn.addEventListener('click', () => fileInput.click());
    
    fileInput.addEventListener('change', handleFileSelect);
    
    uploadZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadZone.style.borderColor = '#2980b9';
        uploadZone.style.backgroundColor = '#e8f4fc';
    });
    
    uploadZone.addEventListener('dragleave', () => {
        uploadZone.style.borderColor = '#3498db';
        uploadZone.style.backgroundColor = '#f8fbfe';
    });
    
    uploadZone.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadZone.style.borderColor = '#3498db';
        uploadZone.style.backgroundColor = '#f8fbfe';
        
        if (e.dataTransfer.files.length) {
            fileInput.files = e.dataTransfer.files;
            handleFileSelect({ target: fileInput });
        }
    });
    
    cancelBtn.addEventListener('click', cancelUpload);
    processBtn.addEventListener('click', processData);
    
    // Tab switching
    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const tabId = btn.getAttribute('data-tab');
            
            // Update active tab button
            tabBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            
            // Show corresponding tab content
            document.querySelectorAll('.tab-content').forEach(content => {
                content.classList.remove('active');
            });
            
            document.getElementById(`${tabId}-tab`).classList.add('active');
        });
    });
    
    // Download buttons
    downloadOriginalBtn.addEventListener('click', () => downloadFile(originalData, 'data_presensi_asli.xlsx'));
    downloadProcessedBtn.addEventListener('click', () => downloadFile(processedData, 'data_presensi_terhitung.xlsx'));
    downloadBothBtn.addEventListener('click', downloadBothFiles);
    
    // File handling function
    function handleFileSelect(event) {
        const file = event.target.files[0];
        if (!file) return;
        
        // Check file extension
        const fileExt = file.name.split('.').pop().toLowerCase();
        if (!['xlsx', 'xls', 'csv'].includes(fileExt)) {
            alert('Format file tidak didukung. Harap unggah file Excel (.xlsx, .xls) atau CSV.');
            return;
        }
        
        currentFile = file;
        
        // Show file info
        fileName.textContent = file.name;
        fileSize.textContent = formatFileSize(file.size);
        fileInfo.style.display = 'block';
        
        // Simulate upload progress
        simulateUploadProgress();
        
        // Read and parse the Excel file
        readExcelFile(file);
    }
    
    // Simulate upload progress
    function simulateUploadProgress() {
        // Clear any existing interval
        if (uploadProgressInterval) {
            clearInterval(uploadProgressInterval);
        }
        
        let progress = 0;
        uploadProgressInterval = setInterval(() => {
            progress += Math.random() * 15;
            if (progress >= 100) {
                progress = 100;
                clearInterval(uploadProgressInterval);
                processBtn.disabled = false;
                processBtn.innerHTML = '<i class="fas fa-calculator"></i> Proses Data Presensi';
            }
            
            progressBar.style.width = `${progress}%`;
            progressText.textContent = `${Math.round(progress)}%`;
        }, 100);
    }
    
    // Read Excel file using SheetJS
    function readExcelFile(file) {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                // Get the first sheet
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                
                // Convert to JSON
                const jsonData = XLSX.utils.sheet_to_json(worksheet);
                
                if (jsonData.length === 0) {
                    throw new Error('File Excel kosong atau tidak memiliki data.');
                }
                
                // Store original data
                originalData = jsonData;
                
                // Display original data in table
                displayOriginalData(jsonData);
                
                showNotification('File berhasil diunggah dan siap diproses.', 'success');
                
            } catch (error) {
                console.error('Error reading Excel file:', error);
                showNotification('Terjadi kesalahan saat membaca file. Pastikan format file benar.', 'error');
                cancelUpload();
            }
        };
        
        reader.onerror = function() {
            showNotification('Gagal membaca file. Coba lagi.', 'error');
            cancelUpload();
        };
        
        reader.readAsArrayBuffer(file);
    }
    
    // Display original data in table
    function displayOriginalData(data) {
        // Clear table
        originalTable.innerHTML = '';
        
        // Limit to first 10 rows for preview
        const previewData = data.slice(0, 10);
        
        previewData.forEach((row, index) => {
            const tr = document.createElement('tr');
            
            // Extract values with various possible column names
            const nama = row.Nama || row.nama || row.NAMA || row.Name || row.name || '-';
            const nik = row.NIK || row.nik || row.Nik || row.ID || row.id || '-';
            const tanggal = row.Tanggal || row.tanggal || row.TANGGAL || row.Date || row.date || '-';
            const jamMasuk = row['Jam Masuk'] || row['jam masuk'] || row.JamMasuk || row.jamMasuk || row['Jam masuk'] || row['Check-in'] || row.checkin || '-';
            const jamKeluar = row['Jam Keluar'] || row['jam keluar'] || row.JamKeluar || row.jamKeluar || row['Jam keluar'] || row['Check-out'] || row.checkout || '-';
            
            tr.innerHTML = `
                <td>${index + 1}</td>
                <td>${nama}</td>
                <td>${nik}</td>
                <td>${formatDate(tanggal)}</td>
                <td>${formatTime(jamMasuk)}</td>
                <td>${formatTime(jamKeluar)}</td>
            `;
            
            originalTable.appendChild(tr);
        });
        
        // If there are more rows, add a note
        if (data.length > 10) {
            const tr = document.createElement('tr');
            tr.innerHTML = `<td colspan="6" style="text-align: center; color: #7f8c8d; font-style: italic;">... dan ${data.length - 10} baris lainnya</td>`;
            originalTable.appendChild(tr);
        }
    }
    
    // Process data and calculate salaries
    function processData() {
        if (originalData.length === 0) {
            alert('Tidak ada data untuk diproses. Silakan unggah file terlebih dahulu.');
            return;
        }
        
        // Show processing indicator
        processBtn.disabled = true;
        processBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Memproses...';
        
        // Get calculation parameters
        const salaryPerHour = parseFloat(document.getElementById('salary-per-hour').value) || 50000;
        const overtimeRate = parseFloat(document.getElementById('overtime-rate').value) || 75000;
        const taxRate = parseFloat(document.getElementById('tax-rate').value) || 5;
        
        // Process data with a slight delay to show processing
        setTimeout(() => {
            try {
                processedData = calculateSalaries(originalData, salaryPerHour, overtimeRate, taxRate);
                
                // Display processed data
                displayProcessedData(processedData);
                
                // Update statistics
                updateStatistics(processedData);
                
                // Show results section
                resultsSection.style.display = 'block';
                
                // Reset process button
                processBtn.disabled = false;
                processBtn.innerHTML = '<i class="fas fa-calculator"></i> Proses Data Presensi';
                
                // Scroll to results
                resultsSection.scrollIntoView({ behavior: 'smooth' });
                
                showNotification('Data berhasil diproses! Lihat hasil di bawah.', 'success');
            } catch (error) {
                console.error('Error processing data:', error);
                showNotification('Terjadi kesalahan saat memproses data. Periksa format data.', 'error');
                
                // Reset process button
                processBtn.disabled = false;
                processBtn.innerHTML = '<i class="fas fa-calculator"></i> Proses Data Presensi';
            }
        }, 800);
    }
    
    // Calculate salaries from attendance data
    function calculateSalaries(data, salaryPerHour, overtimeRate, taxRate) {
        // Group data by employee
        const employees = {};
        
        data.forEach(record => {
            // Extract values with various possible column names
            const name = record.Nama || record.nama || record.NAMA || record.Name || record.name || 'Unknown';
            const nik = record.NIK || record.nik || record.Nik || record.ID || row.id || 'Unknown';
            const date = record.Tanggal || record.tanggal || record.TANGGAL || record.Date || record.date || '';
            const timeIn = record['Jam Masuk'] || record['jam masuk'] || row.JamMasuk || row.jamMasuk || row['Jam masuk'] || row['Check-in'] || row.checkin || '';
            const timeOut = record['Jam Keluar'] || record['jam keluar'] || row.JamKeluar || row.jamKeluar || row['Jam keluar'] || row['Check-out'] || row.checkout || '';
            
            // Create employee key
            const key = `${name}_${nik}`;
            
            if (!employees[key]) {
                employees[key] = {
                    name: name,
                    nik: nik,
                    totalDays: 0,
                    totalHours: 0,
                    overtimeHours: 0,
                    regularHours: 0
                };
            }
            
            // Calculate hours worked
            const hoursWorked = calculateHours(timeIn, timeOut);
            
            if (hoursWorked > 0) {
                employees[key].totalDays += 1;
                employees[key].totalHours += hoursWorked;
                
                // Regular hours (max 8 hours per day)
                const regular = Math.min(hoursWorked, 8);
                employees[key].regularHours += regular;
                
                // Overtime hours (hours beyond 8)
                const overtime = Math.max(hoursWorked - 8, 0);
                employees[key].overtimeHours += overtime;
            }
        });
        
        // Convert to array and calculate salaries
        const result = Object.values(employees).map((emp, index) => {
            const regularSalary = emp.regularHours * salaryPerHour;
            const overtimeSalary = emp.overtimeHours * overtimeRate;
            const grossSalary = regularSalary + overtimeSalary;
            const taxAmount = grossSalary * (taxRate / 100);
            const netSalary = grossSalary - taxAmount;
            
            return {
                No: index + 1,
                Nama: emp.name,
                NIK: emp.nik,
                'Total Hari': emp.totalDays,
                'Total Jam': emp.totalHours.toFixed(2),
                'Lembur (jam)': emp.overtimeHours.toFixed(2),
                'Gaji Pokok': formatCurrency(regularSalary),
                'Uang Lembur': formatCurrency(overtimeSalary),
                'Pajak': formatCurrency(taxAmount),
                'Gaji Bersih': formatCurrency(netSalary),
                // Raw values for calculations and export
                _regularHours: emp.regularHours,
                _overtimeHours: emp.overtimeHours,
                _regularSalary: regularSalary,
                _overtimeSalary: overtimeSalary,
                _taxAmount: taxAmount,
                _netSalary: netSalary
            };
        });
        
        return result;
    }
    
    // Calculate hours between two time strings
    function calculateHours(timeIn, timeOut) {
        if (!timeIn || !timeOut || timeIn === '-' || timeOut === '-') return 0;
        
        try {
            // Convert to string and clean up
            const timeInStr = timeIn.toString().trim();
            const timeOutStr = timeOut.toString().trim();
            
            // Handle various time formats
            let inHours, inMinutes, outHours, outMinutes;
            
            // Try to parse time (e.g., "08:00", "08:00:00", "8.5", "8,5")
            if (timeInStr.includes(':')) {
                const inParts = timeInStr.split(':');
                inHours = parseInt(inParts[0]) || 0;
                inMinutes = parseInt(inParts[1]) || 0;
            } else if (timeInStr.includes('.') || timeInStr.includes(',')) {
                const separator = timeInStr.includes('.') ? '.' : ',';
                const inParts = timeInStr.split(separator);
                inHours = parseInt(inParts[0]) || 0;
                inMinutes = Math.round((parseFloat(timeInStr) - inHours) * 60);
            } else {
                inHours = parseInt(timeInStr) || 0;
                inMinutes = 0;
            }
            
            if (timeOutStr.includes(':')) {
                const outParts = timeOutStr.split(':');
                outHours = parseInt(outParts[0]) || 0;
                outMinutes = parseInt(outParts[1]) || 0;
            } else if (timeOutStr.includes('.') || timeOutStr.includes(',')) {
                const separator = timeOutStr.includes('.') ? '.' : ',';
                const outParts = timeOutStr.split(separator);
                outHours = parseInt(outParts[0]) || 0;
                outMinutes = Math.round((parseFloat(timeOutStr) - outHours) * 60);
            } else {
                outHours = parseInt(timeOutStr) || 0;
                outMinutes = 0;
            }
            
            // Handle cases where out time might be next day (e.g., work overnight)
            if (outHours < inHours) {
                outHours += 24;
            }
            
            // Calculate difference in hours
            let hours = outHours - inHours;
            let minutes = outMinutes - inMinutes;
            
            // Adjust for negative minutes
            if (minutes < 0) {
                hours -= 1;
                minutes += 60;
            }
            
            const totalHours = hours + (minutes / 60);
            
            // Return hours, but ensure it's not negative
            return Math.max(totalHours, 0);
            
        } catch (error) {
            console.error('Error calculating hours:', error, timeIn, timeOut);
            return 0;
        }
    }
    
    // Display processed data in table
    function displayProcessedData(data) {
        // Clear table
        processedTable.innerHTML = '';
        
        // Limit to first 10 rows for preview
        const previewData = data.slice(0, 10);
        
        previewData.forEach(row => {
            const tr = document.createElement('tr');
            
            tr.innerHTML = `
                <td>${row.No}</td>
                <td>${row.Nama}</td>
                <td>${row.NIK}</td>
                <td>${row['Total Hari']}</td>
                <td>${row['Total Jam']}</td>
                <td>${row['Lembur (jam)']}</td>
                <td>${row['Gaji Pokok']}</td>
                <td>${row['Uang Lembur']}</td>
                <td>${row['Pajak']}</td>
                <td>${row['Gaji Bersih']}</td>
            `;
            
            processedTable.appendChild(tr);
        });
        
        // If there are more rows, add a note
        if (data.length > 10) {
            const tr = document.createElement('tr');
            tr.innerHTML = `<td colspan="10" style="text-align: center; color: #7f8c8d; font-style: italic;">... dan ${data.length - 10} baris lainnya</td>`;
            processedTable.appendChild(tr);
        }
    }
    
    // Update statistics display
    function updateStatistics(data) {
        // Calculate totals
        const totalEmp = data.length;
        let totalDaysWorked = 0;
        let totalSalaryPaid = 0;
        let totalOvertimeHours = 0;
        let totalRegularSalary = 0;
        let totalOvertimeSalary = 0;
        let totalTax = 0;
        
        data.forEach(emp => {
            totalDaysWorked += emp['Total Hari'];
            totalSalaryPaid += emp['_netSalary'];
            totalOvertimeHours += emp['_overtimeHours'];
            totalRegularSalary += emp['_regularSalary'];
            totalOvertimeSalary += emp['_overtimeSalary'];
            totalTax += emp['_taxAmount'];
        });
        
        // Update DOM
        totalEmployees.textContent = totalEmp;
        totalDays.textContent = totalDaysWorked;
        totalSalary.textContent = formatCurrency(totalSalaryPaid);
        totalOvertime.textContent = totalOvertimeHours.toFixed(1) + ' jam';
        
        // Log for debugging
        console.log('Statistics updated:', {
            employees: totalEmp,
            days: totalDaysWorked,
            salary: totalSalaryPaid,
            overtime: totalOvertimeHours
        });
    }
    
    // Cancel upload and reset
    function cancelUpload() {
        // Clear upload progress interval
        if (uploadProgressInterval) {
            clearInterval(uploadProgressInterval);
            uploadProgressInterval = null;
        }
        
        fileInput.value = '';
        fileInfo.style.display = 'none';
        processBtn.disabled = true;
        processBtn.innerHTML = '<i class="fas fa-calculator"></i> Proses Data Presensi';
        progressBar.style.width = '0%';
        progressText.textContent = '0%';
        originalData = [];
        originalTable.innerHTML = '';
        
        if (resultsSection.style.display !== 'none') {
            resultsSection.style.display = 'none';
        }
        
        // Clear processed data
        processedData = [];
        processedTable.innerHTML = '';
    }
    
    // Download file
    function downloadFile(data, filename) {
        if (!data || data.length === 0) {
            alert('Tidak ada data untuk diunduh.');
            return;
        }
        
        try {
            // Prepare data for export (remove internal fields starting with _)
            const exportData = data.map(row => {
                const exportRow = {};
                for (const key in row) {
                    if (!key.startsWith('_')) {
                        exportRow[key] = row[key];
                    }
                }
                return exportRow;
            });
            
            // Create worksheet
            const worksheet = XLSX.utils.json_to_sheet(exportData);
            
            // Auto-size columns
            const wscols = [
                {wch: 5},   // No
                {wch: 20},  // Nama
                {wch: 15},  // NIK
                {wch: 10},  // Total Hari
                {wch: 10},  // Total Jam
                {wch: 12},  // Lembur (jam)
                {wch: 15},  // Gaji Pokok
                {wch: 15},  // Uang Lembur
                {wch: 15},  // Pajak
                {wch: 15}   // Gaji Bersih
            ];
            worksheet['!cols'] = wscols;
            
            // Create workbook
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, 'Data Presensi');
            
            // Generate and download file
            XLSX.writeFile(workbook, filename);
            
            // Show success message
            showNotification(`File ${filename} berhasil diunduh.`, 'success');
        } catch (error) {
            console.error('Error downloading file:', error);
            showNotification('Terjadi kesalahan saat mengunduh file.', 'error');
        }
    }
    
    // Download both files as ZIP
    function downloadBothFiles() {
        if (originalData.length === 0 || processedData.length === 0) {
            alert('Tidak ada data untuk diunduh. Silakan proses data terlebih dahulu.');
            return;
        }
        
        try {
            // Download processed data first
            downloadFile(processedData, 'data_presensi_terhitung.xlsx');
            
            // Prepare original data for export
            const exportOriginalData = originalData.map(row => {
                const exportRow = {};
                for (const key in row) {
                    exportRow[key] = row[key];
                }
                return exportRow;
            });
            
            // Wait a bit then download original
            setTimeout(() => {
                // Create worksheet for original data
                const worksheet = XLSX.utils.json_to_sheet(exportOriginalData);
                
                // Create workbook
                const workbook = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(workbook, worksheet, 'Data Presensi');
                
                // Generate and download file
                XLSX.writeFile(workbook, 'data_presensi_asli.xlsx');
                
                showNotification('Kedua file berhasil diunduh.', 'success');
            }, 500);
            
        } catch (error) {
            console.error('Error downloading files:', error);
            showNotification('Terjadi kesalahan saat mengunduh file.', 'error');
        }
    }
    
    // Helper functions
    function formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }
    
    function formatCurrency(amount) {
        return 'Rp ' + amount.toFixed(0).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    }
    
    function formatDate(dateValue) {
        if (!dateValue) return '-';
        
        // If it's already a string, return as is
        if (typeof dateValue === 'string') {
            return dateValue;
        }
        
        // If it's a number (Excel date serial), convert to date
        if (typeof dateValue === 'number') {
            const date = new Date((dateValue - 25569) * 86400 * 1000);
            return date.toLocaleDateString('id-ID');
        }
        
        // If it's a date object
        if (dateValue instanceof Date) {
            return dateValue.toLocaleDateString('id-ID');
        }
        
        return dateValue.toString();
    }
    
    function formatTime(timeValue) {
        if (!timeValue) return '-';
        
        // If it's already a string, return as is
        if (typeof timeValue === 'string') {
            return timeValue;
        }
        
        // If it's a number (Excel time fraction), convert to time
        if (typeof timeValue === 'number') {
            const totalSeconds = Math.round(timeValue * 86400);
            const hours = Math.floor(totalSeconds / 3600);
            const minutes = Math.floor((totalSeconds % 3600) / 60);
            return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
        }
        
        return timeValue.toString();
    }
    
    function showNotification(message, type) {
        // Remove existing notifications
        const existingNotifications = document.querySelectorAll('.notification');
        existingNotifications.forEach(notification => {
            notification.remove();
        });
        
        // Create notification element
        const notification = document.createElement('div');
        notification.className = `notification ${type}`;
        notification.innerHTML = `
            <i class="fas fa-${type === 'success' ? 'check-circle' : 'exclamation-circle'}"></i>
            <span>${message}</span>
        `;
        
        // Add styles
        notification.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            background-color: ${type === 'success' ? '#2ecc71' : '#e74c3c'};
            color: white;
            padding: 15px 20px;
            border-radius: 6px;
            display: flex;
            align-items: center;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            z-index: 1000;
            animation: slideIn 0.3s ease-out;
        `;
        
        // Add icon styles
        notification.querySelector('i').style.marginRight = '10px';
        
        // Add keyframe animation
        if (!document.querySelector('#notification-styles')) {
            const style = document.createElement('style');
            style.id = 'notification-styles';
            style.textContent = `
                @keyframes slideIn {
                    from { transform: translateX(100%); opacity: 0; }
                    to { transform: translateX(0); opacity: 1; }
                }
                @keyframes slideOut {
                    from { transform: translateX(0); opacity: 1; }
                    to { transform: translateX(100%); opacity: 0; }
                }
            `;
            document.head.appendChild(style);
        }
        
        // Add to document
        document.body.appendChild(notification);
        
        // Remove after 3 seconds
        setTimeout(() => {
            notification.style.animation = 'slideOut 0.3s ease-out';
            setTimeout(() => {
                if (notification.parentNode) {
                    notification.parentNode.removeChild(notification);
                }
            }, 300);
        }, 3000);
    }
    
    // Add sample data button for testing
    function addSampleDataButton() {
        const sampleBtn = document.createElement('button');
        sampleBtn.className = 'btn';
        sampleBtn.innerHTML = '<i class="fas fa-vial"></i> Gunakan Data Contoh';
        sampleBtn.style.cssText = `
            position: fixed;
            bottom: 20px;
            right: 20px;
            background-color: #9b59b6;
            color: white;
            z-index: 100;
            padding: 10px 15px;
            font-size: 0.9rem;
        `;
        
        sampleBtn.addEventListener('click', loadSampleData);
        document.body.appendChild(sampleBtn);
    }
    
    // Load sample data for testing
    function loadSampleData() {
        // Create sample data
        const sampleData = [
            { Nama: 'Ahmad Santoso', NIK: 'EMP001', Tanggal: '2023-10-01', 'Jam Masuk': '08:00', 'Jam Keluar': '17:00' },
            { Nama: 'Budi Wijaya', NIK: 'EMP002', Tanggal: '2023-10-01', 'Jam Masuk': '08:30', 'Jam Keluar': '17:30' },
            { Nama: 'Citra Dewi', NIK: 'EMP003', Tanggal: '2023-10-01', 'Jam Masuk': '09:00', 'Jam Keluar': '18:00' },
            { Nama: 'Ahmad Santoso', NIK: 'EMP001', Tanggal: '2023-10-02', 'Jam Masuk': '08:00', 'Jam Keluar': '19:00' },
            { Nama: 'Budi Wijaya', NIK: 'EMP002', Tanggal: '2023-10-02', 'Jam Masuk': '08:15', 'Jam Keluar': '16:45' },
            { Nama: 'Citra Dewi', NIK: 'EMP003', Tanggal: '2023-10-02', 'Jam Masuk': '09:00', 'Jam Keluar': '20:00' },
            { Nama: 'Ahmad Santoso', NIK: 'EMP001', Tanggal: '2023-10-03', 'Jam Masuk': '08:00', 'Jam Keluar': '17:00' },
            { Nama: 'Budi Wijaya', NIK: 'EMP002', Tanggal: '2023-10-03', 'Jam Masuk': '08:30', 'Jam Keluar': '18:30' },
            { Nama: 'Citra Dewi', NIK: 'EMP003', Tanggal: '2023-10-03', 'Jam Masuk': '09:00', 'Jam Keluar': '17:00' },
        ];
        
        // Set as original data
        originalData = sampleData;
        
        // Display in table
        displayOriginalData(sampleData);
        
        // Show file info (simulate upload)
        fileName.textContent = 'contoh_data_presensi.xlsx';
        fileSize.textContent = '2.5 KB';
        fileInfo.style.display = 'block';
        progressBar.style.width = '100%';
        progressText.textContent = '100%';
        processBtn.disabled = false;
        
        showNotification('Data contoh berhasil dimuat. Klik "Proses Data Presensi" untuk menghitung.', 'success');
    }
    
    // Initialize
    addSampleDataButton();
    console.log('Sistem pengolahan data presensi siap digunakan.');
});
