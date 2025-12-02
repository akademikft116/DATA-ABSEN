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
        let progress = 0;
        const interval = setInterval(() => {
            progress += Math.random() * 10;
            if (progress >= 100) {
                progress = 100;
                clearInterval(interval);
                processBtn.disabled = false;
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
                
                // Store original data
                originalData = jsonData;
                
                // Display original data in table
                displayOriginalData(jsonData);
                
            } catch (error) {
                console.error('Error reading Excel file:', error);
                alert('Terjadi kesalahan saat membaca file Excel. Pastikan format file benar.');
            }
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
            
            tr.innerHTML = `
                <td>${index + 1}</td>
                <td>${row.Nama || row.nama || row.NAMA || ''}</td>
                <td>${row.NIK || row.nik || row.Nik || ''}</td>
                <td>${row.Tanggal || row.tanggal || row.TANGGAL || ''}</td>
                <td>${row['Jam Masuk'] || row['jam masuk'] || row.JamMasuk || row.jamMasuk || ''}</td>
                <td>${row['Jam Keluar'] || row['jam keluar'] || row.JamKeluar || row.jamKeluar || ''}</td>
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
        
        // Get calculation parameters
        const salaryPerHour = parseFloat(document.getElementById('salary-per-hour').value) || 50000;
        const overtimeRate = parseFloat(document.getElementById('overtime-rate').value) || 75000;
        const taxRate = parseFloat(document.getElementById('tax-rate').value) || 5;
        
        // Process data
        processedData = calculateSalaries(originalData, salaryPerHour, overtimeRate, taxRate);
        
        // Display processed data
        displayProcessedData(processedData);
        
        // Update statistics
        updateStatistics(processedData);
        
        // Show results section
        resultsSection.style.display = 'block';
        
        // Scroll to results
        resultsSection.scrollIntoView({ behavior: 'smooth' });
    }
    
    // Calculate salaries from attendance data
    function calculateSalaries(data, salaryPerHour, overtimeRate, taxRate) {
        // Group data by employee
        const employees = {};
        
        data.forEach(record => {
            const name = record.Nama || record.nama || record.NAMA || 'Unknown';
            const nik = record.NIK || record.nik || record.Nik || 'Unknown';
            const date = record.Tanggal || record.tanggal || record.TANGGAL || '';
            const timeIn = record['Jam Masuk'] || record['jam masuk'] || record.JamMasuk || record.jamMasuk || '';
            const timeOut = record['Jam Keluar'] || record['jam keluar'] || record.JamKeluar || record.jamKeluar || '';
            
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
        const result = Object.values(employees).map(emp => {
            const regularSalary = emp.regularHours * salaryPerHour;
            const overtimeSalary = emp.overtimeHours * overtimeRate;
            const grossSalary = regularSalary + overtimeSalary;
            const taxAmount = grossSalary * (taxRate / 100);
            const netSalary = grossSalary - taxAmount;
            
            return {
                No: 0, // Will be filled later
                Nama: emp.name,
                NIK: emp.nik,
                'Total Hari': emp.totalDays,
                'Total Jam': emp.totalHours.toFixed(2),
                'Lembur (jam)': emp.overtimeHours.toFixed(2),
                'Gaji Pokok': formatCurrency(regularSalary),
                'Uang Lembur': formatCurrency(overtimeSalary),
                'Pajak': formatCurrency(taxAmount),
                'Gaji Bersih': formatCurrency(netSalary),
                // Raw values for calculations
                _regularHours: emp.regularHours,
                _overtimeHours: emp.overtimeHours,
                _regularSalary: regularSalary,
                _overtimeSalary: overtimeSalary,
                _taxAmount: taxAmount,
                _netSalary: netSalary
            };
        });
        
        // Add numbering
        result.forEach((item, index) => {
            item.No = index + 1;
        });
        
        return result;
    }
    
    // Calculate hours between two time strings
    function calculateHours(timeIn, timeOut) {
        if (!timeIn || !timeOut) return 0;
        
        try {
            // Parse time strings (assuming format like "08:00" or "08:00:00")
            const [inHours, inMinutes] = timeIn.toString().split(':').map(Number);
            const [outHours, outMinutes] = timeOut.toString().split(':').map(Number);
            
            // Calculate difference in hours
            let hours = outHours - inHours;
            let minutes = outMinutes - inMinutes;
            
            // Adjust for negative minutes
            if (minutes < 0) {
                hours -= 1;
                minutes += 60;
            }
            
            return hours + (minutes / 60);
        } catch (error) {
            console.error('Error calculating hours:', error);
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
        
        data.forEach(emp => {
            totalDaysWorked += emp['Total Hari'];
            totalSalaryPaid += emp['_netSalary'];
            totalOvertimeHours += emp['_overtimeHours'];
        });
        
        // Update DOM
        totalEmployees.textContent = totalEmp;
        totalDays.textContent = totalDaysWorked;
        totalSalary.textContent = formatCurrency(totalSalaryPaid);
        totalOvertime.textContent = totalOvertimeHours.toFixed(1) + ' jam';
    }
    
    // Cancel upload and reset
    function cancelUpload() {
        fileInput.value = '';
        fileInfo.style.display = 'none';
        processBtn.disabled = true;
        progressBar.style.width = '0%';
        progressText.textContent = '0%';
        originalData = [];
        originalTable.innerHTML = '';
        
        if (resultsSection.style.display !== 'none') {
            resultsSection.style.display = 'none';
        }
    }
    
    // Download file
    function downloadFile(data, filename) {
        if (!data || data.length === 0) {
            alert('Tidak ada data untuk diunduh.');
            return;
        }
        
        try {
            // Create worksheet
            const worksheet = XLSX.utils.json_to_sheet(data);
            
            // Create workbook
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, 'Data');
            
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
            alert('Tidak ada data untuk diunduh.');
            return;
        }
        
        try {
            // Create workbook for original data
            const originalWorksheet = XLSX.utils.json_to_sheet(originalData);
            const originalWorkbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(originalWorkbook, originalWorksheet, 'Data Presensi');
            
            // Create workbook for processed data
            const processedWorksheet = XLSX.utils.json_to_sheet(processedData);
            const processedWorkbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(processedWorkbook, processedWorksheet, 'Data Terhitung');
            
            // Download first file
            XLSX.writeFile(originalWorkbook, 'data_presensi_asli.xlsx');
            
            // Small delay before downloading second file
            setTimeout(() => {
                XLSX.writeFile(processedWorkbook, 'data_presensi_terhitung.xlsx');
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
    
    function showNotification(message, type) {
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
        const style = document.createElement('style');
        style.textContent = `
            @keyframes slideIn {
                from { transform: translateX(100%); opacity: 0; }
                to { transform: translateX(0); opacity: 1; }
            }
        `;
        document.head.appendChild(style);
        
        // Add to document
        document.body.appendChild(notification);
        
        // Remove after 3 seconds
        setTimeout(() => {
            notification.style.animation = 'slideIn 0.3s ease-out reverse';
            setTimeout(() => {
                if (notification.parentNode) {
                    notification.parentNode.removeChild(notification);
                }
            }, 300);
        }, 3000);
    }
    
    // Initialize with sample data for demonstration
    initializeSampleData();
    
    function initializeSampleData() {
        // This is just for demonstration - in a real app, you would load actual data
        console.log('Sistem presensi siap digunakan.');
    }
});
