// Main JavaScript file - Mengatur seluruh aplikasi

import { processExcelFile, calculateSalaries } from './utils/excelProcessor.js';
import { generateReport, downloadFile } from './utils/reportGenerator.js';
import { formatCurrency, formatDate, formatTime, calculateHours } from './utils/dataFormatter.js';

// Global variables
let originalData = [];
let processedData = [];
let currentFile = null;
let uploadProgressInterval = null;
let hoursChart = null;
let salaryChart = null;

// DOM Elements
const loadingScreen = document.getElementById('loading-screen');
const mainContainer = document.getElementById('main-container');
const excelFileInput = document.getElementById('excel-file');
const uploadArea = document.getElementById('upload-area');
const browseBtn = document.getElementById('browse-btn');
const filePreview = document.getElementById('file-preview');
const processBtn = document.getElementById('process-data');
const resultsSection = document.getElementById('results-section');
const cancelUploadBtn = document.getElementById('cancel-upload');

// Event Listeners
document.addEventListener('DOMContentLoaded', initializeApp);
excelFileInput.addEventListener('change', handleFileSelect);
browseBtn.addEventListener('click', () => excelFileInput.click());
uploadArea.addEventListener('click', () => excelFileInput.click());
processBtn.addEventListener('click', processData);
cancelUploadBtn.addEventListener('click', cancelUpload);

// Initialize application
function initializeApp() {
    // Set current date
    const now = new Date();
    const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
    document.getElementById('current-date').textContent = now.toLocaleDateString('id-ID', options);
    
    // Initialize event listeners for tabs
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            const tabId = this.getAttribute('data-tab');
            switchTab(tabId);
        });
    });
    
    // Initialize help modal
    const helpBtn = document.getElementById('help-btn');
    const closeHelpBtns = document.querySelectorAll('#close-help, #close-help-btn');
    const helpModal = document.getElementById('help-modal');
    
    helpBtn.addEventListener('click', () => helpModal.classList.add('active'));
    closeHelpBtns.forEach(btn => {
        btn.addEventListener('click', () => helpModal.classList.remove('active'));
    });
    
    // Initialize template download
    const templateBtn = document.getElementById('template-btn');
    templateBtn.addEventListener('click', downloadTemplate);
    
    // Initialize reset config
    const resetBtn = document.getElementById('reset-config');
    resetBtn.addEventListener('click', resetConfig);
    
    // Initialize download buttons
    document.getElementById('download-original').addEventListener('click', () => downloadReport('original'));
    document.getElementById('download-processed').addEventListener('click', () => downloadReport('processed'));
    document.getElementById('download-both').addEventListener('click', () => downloadReport('both'));
    
    // Hide loading screen after 2 seconds
    setTimeout(() => {
        loadingScreen.style.opacity = '0';
        setTimeout(() => {
            loadingScreen.style.display = 'none';
            mainContainer.classList.add('loaded');
        }, 500);
    }, 2000);
}

// Handle file selection
async function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    currentFile = file;
    
    // Show file preview
    showFilePreview(file);
    
    // Start upload simulation
    simulateUploadProgress();
    
    // Process the file
    try {
        const data = await processExcelFile(file);
        originalData = data;
        
        // Update stats
        updateSidebarStats(data);
        
        // Show success message
        showNotification('File berhasil diunggah!', 'success');
        
        // Enable process button
        processBtn.disabled = false;
        
    } catch (error) {
        console.error('Error processing file:', error);
        showNotification('Gagal memproses file. Pastikan format sesuai.', 'error');
        cancelUpload();
    }
}

// Show file preview
function showFilePreview(file) {
    const fileName = document.getElementById('file-name');
    const fileSize = document.getElementById('file-size');
    const fileDate = document.getElementById('file-date');
    
    fileName.textContent = file.name;
    fileSize.textContent = formatFileSize(file.size);
    fileDate.textContent = new Date(file.lastModified).toLocaleDateString('id-ID');
    
    filePreview.style.display = 'block';
    uploadArea.style.display = 'none';
}

// Simulate upload progress
function simulateUploadProgress() {
    if (uploadProgressInterval) clearInterval(uploadProgressInterval);
    
    const progressBar = document.getElementById('upload-progress');
    const progressText = document.getElementById('progress-text');
    
    let progress = 0;
    uploadProgressInterval = setInterval(() => {
        progress += Math.random() * 15;
        if (progress >= 100) {
            progress = 100;
            clearInterval(uploadProgressInterval);
        }
        
        progressBar.style.width = `${progress}%`;
        progressText.textContent = `${Math.round(progress)}%`;
    }, 100);
}

// Process data
function processData() {
    if (originalData.length === 0) {
        showNotification('Tidak ada data untuk diproses.', 'warning');
        return;
    }
    
    // Get configuration values
    const salaryPerHour = parseFloat(document.getElementById('salary-per-hour').value) || 50000;
    const overtimeRate = parseFloat(document.getElementById('overtime-rate').value) || 75000;
    const taxRate = parseFloat(document.getElementById('tax-rate').value) || 5;
    const workHours = parseFloat(document.getElementById('work-hours').value) || 8;
    
    // Show processing state
    processBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Memproses...';
    processBtn.disabled = true;
    
    // Process data with delay to show loading
    setTimeout(() => {
        try {
            // Calculate salaries
            processedData = calculateSalaries(originalData, salaryPerHour, overtimeRate, taxRate, workHours);
            
            // Display results
            displayResults(processedData);
            
            // Show charts
            createCharts(processedData);
            
            // Show results section
            resultsSection.style.display = 'block';
            
            // Scroll to results
            resultsSection.scrollIntoView({ behavior: 'smooth' });
            
            // Show success message
            showNotification('Data berhasil diproses!', 'success');
            
        } catch (error) {
            console.error('Error processing data:', error);
            showNotification('Terjadi kesalahan saat memproses data.', 'error');
        } finally {
            // Reset process button
            processBtn.innerHTML = '<i class="fas fa-calculator"></i> Proses Data';
            processBtn.disabled = false;
        }
    }, 1500);
}

// Display results
function displayResults(data) {
    // Update main statistics
    updateMainStatistics(data);
    
    // Display processed data table
    displayProcessedTable(data);
    
    // Display original data table
    displayOriginalTable();
    
    // Display summaries
    displaySummaries(data);
}

// Update main statistics
function updateMainStatistics(data) {
    const totalKaryawan = new Set(data.map(item => item.nama)).size;
    const totalHari = data.reduce((sum, item) => sum + item.totalHari, 0);
    const totalGaji = data.reduce((sum, item) => sum + item.gajiBersih, 0);
    const totalLembur = data.reduce((sum, item) => sum + item.jamLembur, 0);
    
    document.getElementById('total-karyawan').textContent = totalKaryawan;
    document.getElementById('total-hari').textContent = totalHari;
    document.getElementById('total-gaji').textContent = formatCurrency(totalGaji);
    document.getElementById('total-lembur').textContent = totalLembur.toFixed(1) + ' jam';
}

// Display processed table
function displayProcessedTable(data) {
    const tbody = document.getElementById('processed-table-body');
    tbody.innerHTML = '';
    
    data.forEach((item, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${index + 1}</td>
            <td><strong>${item.nama}</strong></td>
            <td>${item.totalHari}</td>
            <td>${item.totalJam.toFixed(2)}</td>
            <td>${item.jamLembur.toFixed(2)}</td>
            <td>${formatCurrency(item.gajiPokok)}</td>
            <td>${formatCurrency(item.uangLembur)}</td>
            <td>${formatCurrency(item.pajak)}</td>
            <td><strong>${formatCurrency(item.gajiBersih)}</strong></td>
        `;
        tbody.appendChild(row);
    });
}

// Display original table
function displayOriginalTable() {
    const tbody = document.getElementById('original-table-body');
    tbody.innerHTML = '';
    
    // Show first 10 records
    const previewData = originalData.slice(0, 10);
    
    previewData.forEach((item, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${index + 1}</td>
            <td>${item.nama || '-'}</td>
            <td>${formatDate(item.tanggal) || '-'}</td>
            <td>${formatTime(item.jamMasuk) || '-'}</td>
            <td>${formatTime(item.jamKeluar) || '-'}</td>
            <td>${item.durasi ? item.durasi.toFixed(2) + ' jam' : '-'}</td>
        `;
        tbody.appendChild(row);
    });
}

// Display summaries
function displaySummaries(data) {
    // Employee summary
    const employeeSummary = document.getElementById('employee-summary');
    const uniqueEmployees = [...new Set(data.map(item => item.nama))];
    
    let employeeHtml = '';
    uniqueEmployees.forEach(employee => {
        const employeeData = data.find(item => item.nama === employee);
        employeeHtml += `
            <div>
                <strong>${employee}</strong><br>
                <small>Total Hari: ${employeeData.totalHari} | Total Jam: ${employeeData.totalJam.toFixed(2)} | Lembur: ${employeeData.jamLembur.toFixed(2)} jam</small>
            </div>
        `;
    });
    employeeSummary.innerHTML = employeeHtml;
    
    // Financial summary
    const financialSummary = document.getElementById('financial-summary');
    const totalGajiPokok = data.reduce((sum, item) => sum + item.gajiPokok, 0);
    const totalUangLembur = data.reduce((sum, item) => sum + item.uangLembur, 0);
    const totalPajak = data.reduce((sum, item) => sum + item.pajak, 0);
    const totalGajiBersih = data.reduce((sum, item) => sum + item.gajiBersih, 0);
    
    financialSummary.innerHTML = `
        <div>Total Gaji Pokok: <strong>${formatCurrency(totalGajiPokok)}</strong></div>
        <div>Total Uang Lembur: <strong>${formatCurrency(totalUangLembur)}</strong></div>
        <div>Total Potongan Pajak: <strong>${formatCurrency(totalPajak)}</strong></div>
        <div style="border-top: 2px solid #3498db; padding-top: 0.5rem; margin-top: 0.5rem;">
            Total Gaji Bersih: <strong style="color: #2ecc71; font-size: 1.1em;">${formatCurrency(totalGajiBersih)}</strong>
        </div>
    `;
}

// Create charts
function createCharts(data) {
    // Destroy existing charts
    if (hoursChart) hoursChart.destroy();
    if (salaryChart) salaryChart.destroy();
    
    // Prepare data for charts
    const employeeNames = [...new Set(data.map(item => item.nama))];
    const totalHours = employeeNames.map(name => {
        const employeeData = data.find(item => item.nama === name);
        return employeeData ? employeeData.totalJam : 0;
    });
    
    const regularHours = employeeNames.map(name => {
        const employeeData = data.find(item => item.nama === name);
        return employeeData ? employeeData.totalJam - employeeData.jamLembur : 0;
    });
    
    const overtimeHours = employeeNames.map(name => {
        const employeeData = data.find(item => item.nama === name);
        return employeeData ? employeeData.jamLembur : 0;
    });
    
    // Hours chart
    const hoursCtx = document.getElementById('hoursChart').getContext('2d');
    hoursChart = new Chart(hoursCtx, {
        type: 'bar',
        data: {
            labels: employeeNames,
            datasets: [
                {
                    label: 'Jam Normal',
                    data: regularHours,
                    backgroundColor: 'rgba(52, 152, 219, 0.7)',
                    borderColor: 'rgba(52, 152, 219, 1)',
                    borderWidth: 1
                },
                {
                    label: 'Jam Lembur',
                    data: overtimeHours,
                    backgroundColor: 'rgba(231, 76, 60, 0.7)',
                    borderColor: 'rgba(231, 76, 60, 1)',
                    borderWidth: 1
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'Jam Kerja'
                    }
                },
                x: {
                    title: {
                        display: true,
                        text: 'Karyawan'
                    }
                }
            }
        }
    });
    
    // Salary chart
    const salaryCtx = document.getElementById('salaryChart').getContext('2d');
    const salaries = employeeNames.map(name => {
        const employeeData = data.find(item => item.nama === name);
        return employeeData ? employeeData.gajiBersih : 0;
    });
    
    salaryChart = new Chart(salaryCtx, {
        type: 'pie',
        data: {
            labels: employeeNames,
            datasets: [{
                data: salaries,
                backgroundColor: [
                    '#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0',
                    '#9966FF', '#FF9F40', '#C9CBCF', '#7E57C2',
                    '#42A5F5', '#26C6DA', '#66BB6A', '#FFCA28'
                ]
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'right',
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const value = context.raw;
                            const percentage = ((value / salaries.reduce((a, b) => a + b, 0)) * 100).toFixed(1);
                            return `${context.label}: ${formatCurrency(value)} (${percentage}%)`;
                        }
                    }
                }
            }
        }
    });
}

// Switch tabs
function switchTab(tabId) {
    // Update active tab button
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('active');
        if (btn.getAttribute('data-tab') === tabId) {
            btn.classList.add('active');
        }
    });
    
    // Show active tab content
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.remove('active');
        if (content.id === `${tabId}-tab`) {
            content.classList.add('active');
        }
    });
}

// Cancel upload
function cancelUpload() {
    if (uploadProgressInterval) {
        clearInterval(uploadProgressInterval);
        uploadProgressInterval = null;
    }
    
    excelFileInput.value = '';
    filePreview.style.display = 'none';
    uploadArea.style.display = 'block';
    processBtn.disabled = true;
    
    // Hide results if shown
    resultsSection.style.display = 'none';
}

// Reset configuration
function resetConfig() {
    document.getElementById('salary-per-hour').value = '';
    document.getElementById('overtime-rate').value = '';
    document.getElementById('tax-rate').value = '';
    document.getElementById('work-hours').value = '8';
    
    showNotification('Konfigurasi telah direset ke nilai default.', 'info');
}

// Download template
function downloadTemplate() {
    // Create sample template
    const templateData = [
        {
            'Nama': 'Contoh Karyawan 1',
            'Tanggal': '2025-11-01',
            'Jam Masuk': '08:00',
            'Jam Keluar': '17:00'
        },
        {
            'Nama': 'Contoh Karyawan 2',
            'Tanggal': '2025-11-01',
            'Jam Masuk': '08:30',
            'Jam Keluar': '17:30'
        },
        {
            'Nama': 'Contoh Karyawan 1',
            'Tanggal': '2025-11-02',
            'Jam Masuk': '08:15',
            'Jam Keluar': '18:15'
        }
    ];
    
    downloadFile(templateData, 'template_data_presensi.xlsx', 'Template Presensi');
    showNotification('Template berhasil diunduh.', 'success');
}

// Download report
async function downloadReport(type) {
    if (type === 'original' && originalData.length === 0) {
        showNotification('Tidak ada data asli untuk diunduh.', 'warning');
        return;
    }
    
    if ((type === 'processed' || type === 'both') && processedData.length === 0) {
        showNotification('Data belum diproses.', 'warning');
        return;
    }
    
    try {
        if (type === 'original') {
            await generateReport(originalData, 'data_presensi_asli.xlsx', 'Data Asli');
            showNotification('Data asli berhasil diunduh.', 'success');
        } else if (type === 'processed') {
            await generateReport(processedData, 'data_presensi_terhitung.xlsx', 'Data Terhitung');
            showNotification('Data terhitung berhasil diunduh.', 'success');
        } else if (type === 'both') {
            await generateReport(originalData, 'data_presensi_asli.xlsx', 'Data Asli');
            setTimeout(async () => {
                await generateReport(processedData, 'data_presensi_terhitung.xlsx', 'Data Terhitung');
                showNotification('Kedua file berhasil diunduh.', 'success');
            }, 500);
        }
    } catch (error) {
        console.error('Error downloading report:', error);
        showNotification('Gagal mengunduh laporan.', 'error');
    }
}

// Update sidebar stats
function updateSidebarStats(data) {
    const uniqueEmployees = new Set(data.map(item => item.nama)).size;
    const totalDays = new Set(data.map(item => item.tanggal)).size;
    
    document.getElementById('stat-employees').textContent = uniqueEmployees;
    document.getElementById('stat-attendance').textContent = totalDays;
}

// Format file size
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// Show notification
function showNotification(message, type = 'info') {
    // Create notification element
    const notification = document.createElement('div');
    notification.className = `notification-${type}`;
    notification.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        padding: 1rem 1.5rem;
        background: ${type === 'success' ? '#2ecc71' : type === 'error' ? '#e74c3c' : '#3498db'};
        color: white;
        border-radius: var(--border-radius-sm);
        box-shadow: var(--shadow-medium);
        z-index: 1000;
        animation: slideInRight 0.3s ease;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    `;
    
    const icon = type === 'success' ? 'check-circle' : type === 'error' ? 'exclamation-circle' : 'info-circle';
    notification.innerHTML = `<i class="fas fa-${icon}"></i> ${message}`;
    
    document.body.appendChild(notification);
    
    // Remove notification after 3 seconds
    setTimeout(() => {
        notification.style.animation = 'slideOutRight 0.3s ease';
        setTimeout(() => {
            if (notification.parentNode) {
                notification.parentNode.removeChild(notification);
            }
        }, 300);
    }, 3000);
}

// Add CSS animations for notifications
const style = document.createElement('style');
style.textContent = `
    @keyframes slideInRight {
        from {
            transform: translateX(100%);
            opacity: 0;
        }
        to {
            transform: translateX(0);
            opacity: 1;
        }
    }
    
    @keyframes slideOutRight {
        from {
            transform: translateX(0);
            opacity: 1;
        }
        to {
            transform: translateX(100%);
            opacity: 0;
        }
    }
`;
document.head.appendChild(style);

// Export for module usage
export { showNotification, formatFileSize };
